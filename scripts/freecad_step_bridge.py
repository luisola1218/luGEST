from __future__ import annotations

import json
import sys
from pathlib import Path


def _dominant_axis(vector: tuple[float, float, float]) -> int:
    values = [abs(float(part or 0.0)) for part in vector]
    return max(range(3), key=lambda index: values[index])


def _classify_wire(wire) -> str:
    curve_kinds = {type(edge.Curve).__name__ for edge in list(wire.Edges or []) if getattr(edge, "Curve", None) is not None}
    if not curve_kinds:
        return "cut"
    if curve_kinds == {"Circle"}:
        bbox = wire.BoundBox
        width = max(float(bbox.XLength), float(bbox.YLength), float(bbox.ZLength))
        ordered = sorted([float(bbox.XLength), float(bbox.YLength), float(bbox.ZLength)], reverse=True)
        height = ordered[1] if len(ordered) > 1 else 0.0
        length = _wire_length_mm(wire)
        expected_circle_length = 3.141592653589793 * max(0.0, (width + height) / 2.0)
        is_round = width > 0.0 and height > 0.0 and abs(width - height) <= max(0.25, width * 0.08)
        is_full_circle = expected_circle_length > 0.0 and abs(length - expected_circle_length) <= max(0.75, expected_circle_length * 0.18)
        if is_round and is_full_circle:
            return "hole"
        return "slot"
    return "slot"


def _wire_length_mm(wire) -> float:
    try:
        return max(0.0, float(getattr(wire, "Length", 0.0) or 0.0))
    except Exception:
        return 0.0


def _wire_signature(wire, plane_axis_index: int, side_sign: int, loop_kind: str) -> tuple[object, ...]:
    bbox = wire.BoundBox
    mins = (float(bbox.XMin), float(bbox.YMin), float(bbox.ZMin))
    maxs = (float(bbox.XMax), float(bbox.YMax), float(bbox.ZMax))
    axes = [index for index in range(3) if index != int(plane_axis_index)]
    return (
        int(plane_axis_index),
        int(side_sign),
        loop_kind,
        round((mins[axes[0]] + maxs[axes[0]]) / 2.0, 3),
        round((mins[axes[1]] + maxs[axes[1]]) / 2.0, 3),
        round(maxs[axes[0]] - mins[axes[0]], 3),
        round(maxs[axes[1]] - mins[axes[1]], 3),
    )


def _round_clean(value: float, digits: int = 3) -> float:
    rounded = round(float(value or 0.0), digits)
    if abs(rounded - round(rounded)) <= 0.001:
        return float(round(rounded))
    return rounded


def _profile_geometry_guess(
    *,
    spans: tuple[float, float, float],
    principal_axis: int,
    plane_faces: list[dict[str, object]],
) -> dict[str, object]:
    cross_axes = [axis for axis in range(3) if axis != int(principal_axis)]
    cross_spans = [float(spans[axis] if axis < len(spans) else 0.0) for axis in cross_axes]
    length_mm = float(spans[principal_axis] if principal_axis < len(spans) else 0.0)
    positions_by_axis: dict[int, list[float]] = {}
    for axis in cross_axes:
        raw_positions = sorted(float(row["position"]) for row in plane_faces if int(row["axis_index"]) == axis)
        cleaned: list[float] = []
        for value in raw_positions:
            if not cleaned or abs(value - cleaned[-1]) > 0.35:
                cleaned.append(value)
        positions_by_axis[axis] = cleaned

    thickness_candidates: list[float] = []
    internal_position_count = 0
    for axis, positions in positions_by_axis.items():
        if len(positions) < 3:
            continue
        span = max(0.0, max(positions) - min(positions))
        if span <= 0.0:
            continue
        for pos in positions[1:-1]:
            if abs(pos - min(positions)) > 0.6 and abs(pos - max(positions)) > 0.6:
                internal_position_count += 1
        diffs = [
            positions[index + 1] - positions[index]
            for index in range(len(positions) - 1)
            if positions[index + 1] - positions[index] > 0.5
        ]
        for diff in diffs:
            if diff <= max(25.0, span * 0.4):
                thickness_candidates.append(float(diff))

    thickness_mm = min(thickness_candidates) if thickness_candidates else 0.0
    leg_a = max(cross_spans) if cross_spans else 0.0
    leg_b = min(cross_spans) if len(cross_spans) > 1 else 0.0
    family_guess = ""
    if thickness_mm > 0.0 and leg_a > 0.0 and leg_b > 0.0:
        if internal_position_count >= 4:
            family_guess = "Tubo"
        elif internal_position_count >= 2:
            family_guess = "Cantoneira"
    section_label = ""
    if family_guess in {"Cantoneira", "Tubo"}:
        section_label = f"{_round_clean(leg_a):g}x{_round_clean(leg_b):g}x{_round_clean(thickness_mm):g}"
    return {
        "profile_length_m": round(length_mm / 1000.0, 4) if length_mm > 0.0 else 0.0,
        "profile_length_mm": round(length_mm, 3),
        "section_spans_mm": [round(value, 3) for value in cross_spans],
        "section_label": section_label,
        "thickness_mm_guess": round(thickness_mm, 3),
        "family_guess": family_guess,
        "internal_plane_count": int(internal_position_count),
    }


def analyze_step_profile(source_path: Path) -> dict[str, object]:
    import Part

    shape = Part.Shape()
    shape.read(str(source_path))
    if shape.isNull():
        return {}

    bound_box = shape.BoundBox
    spans = (float(bound_box.XLength), float(bound_box.YLength), float(bound_box.ZLength))
    principal_axis = max(range(3), key=lambda index: spans[index])
    principal_axis_name = ("x", "y", "z")[principal_axis]
    principal_center = (
        (float(bound_box.XMin) + float(bound_box.XMax)) / 2.0,
        (float(bound_box.YMin) + float(bound_box.YMax)) / 2.0,
        (float(bound_box.ZMin) + float(bound_box.ZMax)) / 2.0,
    )[principal_axis]

    plane_faces: list[dict[str, object]] = []
    for face_index, face in enumerate(list(shape.Faces or []), start=1):
        surface = getattr(face, "Surface", None)
        if type(surface).__name__ != "Plane":
            continue
        axis = getattr(surface, "Axis", None)
        position = getattr(surface, "Position", None)
        if axis is None or position is None:
            continue
        normal = (float(axis.x), float(axis.y), float(axis.z))
        axis_index = _dominant_axis(normal)
        origin = (float(position.x), float(position.y), float(position.z))
        plane_faces.append(
            {
                "index": face_index,
                "face": face,
                "axis_index": axis_index,
                "position": origin[axis_index],
                "area": float(face.Area),
                "wires": len(list(face.Wires or [])),
            }
        )

    end_groups: dict[int, list[dict[str, object]]] = {-1: [], 1: []}
    lateral_groups: dict[tuple[int, int], list[dict[str, object]]] = {}
    for row in plane_faces:
        side_sign = 1 if float(row["position"]) >= principal_center else -1
        row["side_sign"] = side_sign
        if int(row["axis_index"]) == principal_axis:
            end_groups[side_sign].append(row)
        elif int(row["wires"]) > 1:
            key = (int(row["axis_index"]), int(side_sign))
            lateral_groups.setdefault(key, []).append(row)

    lateral_faces: list[dict[str, object]] = []
    for group_rows in lateral_groups.values():
        lateral_faces.append(max(group_rows, key=lambda item: abs(float(item["position"]) - principal_center)))

    selected_end_faces: list[dict[str, object]] = []
    for side_sign in (-1, 1):
        candidates = list(end_groups.get(side_sign, []) or [])
        if not candidates:
            continue
        selected_end_faces.append(max(candidates, key=lambda item: (float(item["area"]), abs(float(item["position"]) - principal_center))))

    seen_features: set[tuple[object, ...]] = set()
    hole_count = 0
    slot_count = 0
    generic_cut_count = 0
    hole_cut_length_mm = 0.0
    slot_cut_length_mm = 0.0
    generic_cut_length_mm = 0.0
    feature_rows: list[dict[str, object]] = []

    for row in lateral_faces:
        face = row["face"]
        outer_wire = getattr(face, "OuterWire", None)
        for wire in list(face.Wires or []):
            if outer_wire is not None and wire.isSame(outer_wire):
                continue
            loop_kind = _classify_wire(wire)
            signature = _wire_signature(wire, int(row["axis_index"]), int(row["side_sign"]), loop_kind)
            if signature in seen_features:
                continue
            seen_features.add(signature)
            cut_length_mm = _wire_length_mm(wire)
            bbox = wire.BoundBox
            feature_rows.append(
                {
                    "face_index": int(row["index"]),
                    "kind": loop_kind,
                    "length_mm": round(cut_length_mm, 3),
                    "signature": list(signature),
                    "bbox": [
                        round(float(bbox.XMin), 3),
                        round(float(bbox.YMin), 3),
                        round(float(bbox.ZMin), 3),
                        round(float(bbox.XMax), 3),
                        round(float(bbox.YMax), 3),
                        round(float(bbox.ZMax), 3),
                    ],
                }
            )
            if loop_kind == "hole":
                hole_count += 1
                hole_cut_length_mm += cut_length_mm
            elif loop_kind == "slot":
                slot_count += 1
                slot_cut_length_mm += cut_length_mm
            else:
                generic_cut_count += 1
                generic_cut_length_mm += cut_length_mm

    side_feature_count = hole_count + slot_count + generic_cut_count
    end_cut_count = len(selected_end_faces)
    end_cut_length_mm = 0.0
    for row in selected_end_faces:
        face = row["face"]
        outer_wire = getattr(face, "OuterWire", None)
        if outer_wire is not None:
            end_cut_length_mm += _wire_length_mm(outer_wire)
    cuts = int(side_feature_count + end_cut_count)
    outer_cuts = int(generic_cut_count + end_cut_count)
    feature_cut_length_mm = hole_cut_length_mm + slot_cut_length_mm + generic_cut_length_mm
    cut_length_mm = feature_cut_length_mm + end_cut_length_mm
    notes = [
        f"analise FreeCAD no eixo principal {principal_axis_name}",
        f"eventos adicionais de corte no perfil: {side_feature_count}",
        f"cortes terminais base: {end_cut_count}",
        f"eventos laser totais: {cuts}",
        f"comprimento de corte medido: {round(cut_length_mm / 1000.0, 3)} m",
    ]
    profile_guess = _profile_geometry_guess(spans=spans, principal_axis=principal_axis, plane_faces=plane_faces)
    if profile_guess.get("family_guess"):
        notes.append(
            f"perfil sugerido: {profile_guess.get('family_guess')} {profile_guess.get('section_label') or ''}".strip()
        )
    if float(profile_guess.get("thickness_mm_guess", 0.0) or 0.0) > 0.0:
        notes.append(f"espessura sugerida: {float(profile_guess.get('thickness_mm_guess', 0.0) or 0.0):g} mm")
    if hole_count:
        notes.append(f"furos circulares detetados: {hole_count}")
    if slot_count:
        notes.append(f"rasgos/contornos nao circulares detetados: {slot_count}")
    if generic_cut_count:
        notes.append(f"aberturas/cortes genericos detetados: {generic_cut_count}")

    return {
        "mode": "freecad_topology",
        "engine": "freecad",
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
        "cut_length_mm": round(cut_length_mm, 3),
        "cut_length_m": round(cut_length_mm / 1000.0, 4),
        "feature_cut_length_mm": round(feature_cut_length_mm, 3),
        "feature_cut_length_m": round(feature_cut_length_mm / 1000.0, 4),
        "end_cut_length_mm": round(end_cut_length_mm, 3),
        "end_cut_length_m": round(end_cut_length_mm / 1000.0, 4),
        "hole_cut_length_mm": round(hole_cut_length_mm, 3),
        "slot_cut_length_mm": round(slot_cut_length_mm, 3),
        "generic_cut_length_mm": round(generic_cut_length_mm, 3),
        "selected_lateral_faces": [int(row["index"]) for row in lateral_faces],
        "selected_end_faces": [int(row["index"]) for row in selected_end_faces],
        "feature_rows": feature_rows,
        "bbox_mm": {
            "x": round(float(spans[0]), 3),
            "y": round(float(spans[1]), 3),
            "z": round(float(spans[2]), 3),
        },
        **profile_guess,
        "note": ". ".join(notes),
    }


def preview_step_profile(source_path: Path) -> dict[str, object]:
    import Part

    shape = Part.Shape()
    shape.read(str(source_path))
    if shape.isNull():
        return {"available": False, "note": "Shape FreeCAD vazia."}

    def project_point(point) -> tuple[float, float, float]:
        x = float(point.x)
        y = float(point.y)
        z = float(point.z)
        return ((x - y) * 0.8660254, z - ((x + y) * 0.5), x + y + z)

    polylines: list[dict[str, object]] = []
    for edge in list(shape.Edges or []):
        try:
            if type(edge.Curve).__name__ == "Line":
                sampled = [vertex.Point for vertex in list(edge.Vertexes or [])]
            else:
                target_points = max(12, min(48, int(round(float(edge.Length) / 12.0))))
                sampled = list(edge.discretize(target_points))
        except Exception:
            sampled = [vertex.Point for vertex in list(edge.Vertexes or [])]
        if len(sampled) < 2:
            continue
        projected = [project_point(point) for point in sampled]
        polylines.append(
            {
                "points": [[round(item[0], 4), round(item[1], 4)] for item in projected],
                "depth": round(sum(item[2] for item in projected) / len(projected), 4),
            }
        )
    return {"available": bool(polylines), "engine": "freecad", "polylines": polylines}


def main() -> int:
    args = list(sys.argv[1:])
    if args and str(args[0]).strip().lower().endswith(".py"):
        args = args[1:]
    if args and str(args[0]).strip() == "--pass":
        args = args[1:]
    if len(args) < 3:
        raise SystemExit(2)
    mode = str(args[0] or "").strip().lower()
    source_path = Path(args[1]).expanduser()
    output_path = Path(args[2]).expanduser()
    payload: dict[str, object]
    try:
        if mode == "analyze":
            payload = analyze_step_profile(source_path)
        elif mode == "preview":
            payload = preview_step_profile(source_path)
        else:
            payload = {"available": False, "note": f"Modo FreeCAD nao suportado: {mode}"}
    except Exception as exc:
        payload = {"available": False, "note": str(exc)}
    output_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return 0
raise SystemExit(main())
